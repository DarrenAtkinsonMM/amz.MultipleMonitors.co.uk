<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<% 
dim pIdProduct

pcSCID=getUserInput(request.QueryString("id"),0)
if not validNum(pcSCID) then
   response.redirect "default.asp"
end if

	query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveDesc FROM pcSales_Completed WHERE pcSales_Completed.pcSC_ID=" & pcSCID & ";"
	set rsS=Server.CreateObject("ADODB.Recordset")
	set rsS=conntemp.execute(query)
					
			if not rsS.eof then
				pcSCID=rsS("pcSC_ID")
				pcSCName=rsS("pcSC_SaveName")
				pcSCDesc=rsS("pcSC_SaveDesc")
			end if
			set rsS=nothing
%> 
<script type=text/javascript>
	function WinResize()
	{
	var showScroll=0;
		if (/Firefox[\/\s](\d+\.\d+)/.test(navigator.userAgent)){
			wH=document.body.scrollHeight+100;
			wW=document.body.scrollWidth+20;
		}
			else
		{
			wH=document.body.scrollHeight+80;
			wW=document.body.scrollWidth+20;
		}
	if (wH>550)
	{
		showScroll=1;
		wH=550;
	}
	if (wW>650)
	{
		showScroll=1;
		wW=650;
	}
	
	window.resizeTo(wW,wH);
	if (showScroll==1) document.body.scroll="yes";
		
	}
</script>
<div id="sm_showdetails">
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
		<h3 class="modal-title"><%=dictLanguage.Item(Session("language")&"_Sale_1") & pcSCName%></h3>
	</div>

		
	<div class="modal-body">
		<strong><%=dictLanguage.Item(Session("language")&"_Sale_2")%></strong>
		<div class="pcSpacer"></div>
		<div class="pcFormItem"><%= pcf_FixHTMLContentPaths(pcSCDesc) %></div>
	</div>

	<div class="modal-footer">
		<button class="btn btn-default" data-dismiss="modal">Close Window</button>
	</div>
</div>

<%
call closeDb()
%>
