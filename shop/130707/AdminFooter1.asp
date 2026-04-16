<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, all of its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit http://www.productcart.com.
%>
        </div>
    <% If pcv_strDisplayType <> "1" Then %>
        <div id="pcCPmainRight">
		</div>>         
    <% End If %>
	</div>        
    <div id="pcFooter">
        <a href="about_terms.asp"><div style="float: left"><img src="images/pc_logo_100.gif" width="100" height="30" alt="ProductCart shopping cart software" border="0" /></div>Use of this software indicates acceptance of the End User License Agreement</a><br /><a href="http://www.productcart.com">Copyright&copy; 2001-<%=Year(now)%> NetSource Commerce. All Rights Reserved. ProductCart&reg; is a registered trademark of NetSource Commerce</a>.
    </div>

	<script language="JavaScript" type="text/javascript">
	<%
	tmpStr=""
	IF lcase(section)<>"quickbooks" AND lcase(section)<>"ebay" AND lcase(pageTitle)<>"productcart ebay add-on" THEN
 
	if session("admin")<>"0" and session("admin")<>"" then
	tmpStr=tmpStr & "$( ""#cp1"" ).accordion( ""option"", ""active"", 0 );"
	tmpStr=tmpStr & "$( ""#cp3"" ).accordion( ""option"", ""active"", 0 );"
	tmpStr=tmpStr & "$( ""#cp4"" ).accordion( ""option"", ""active"", 0 );"
	%>
	$( "#cp1" ).accordion({collapsible: true, header: "h5", active:false});
	$( "#cp1 span").removeClass('ui-icon');
	$( "#cp3" ).accordion({collapsible: true, header: "h5", active:false});
	$( "#cp3 span").removeClass('ui-icon');
	$( "#cp4" ).accordion({collapsible: true, header: "h5", active:false});
	$( "#cp4 span").removeClass('ui-icon');
	<%end if
	END IF%>
	<% 
	if pcv_ShowSmallRecentProducts=1 then 
	tmpStr=tmpStr & "$( ""#cp2"" ).accordion( ""option"", ""active"", 0 );"%>
	$( "#cp2" ).accordion({collapsible: true, header: "h5", active:false});
	$( "#cp2 span").removeClass('ui-icon');
	<% 
	end if 
	%>
	<%
	if pcInt_ShowOrderLegend = 1 then
	tmpStr=tmpStr & "$( ""#cp5"" ).accordion( ""option"", ""active"", 0 );"
	tmpStr=tmpStr & "$( ""#cp6"" ).accordion( ""option"", ""active"", 0 );"
	%>
	$( "#cp5" ).accordion({collapsible: true, header: "h5", active:false});
	$( "#cp5 span").removeClass('ui-icon');
	$( "#cp6" ).accordion({collapsible: true, header: "h5", active:false});
	$( "#cp6 span").removeClass('ui-icon');
	<%
	end if
	%>
	<%=tmpStr%>
</script>

</body>
</html>